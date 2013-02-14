VERSION 5.00
Object = "{AE19B085-7851-4724-8240-EC49EA45E455}#3.0#0"; "pbmasone1.ocx"
Begin VB.Form splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PBMasOne.MasOnePB pb1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4800
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   450
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   360
      Top             =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "loading ....."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   100
      TabIndex        =   0
      Top             =   4470
      Width           =   7575
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   0
      Picture         =   "splash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crColor As Long, ByVal nAlpha As Byte, ByVal dwFlags As Long) As Long
Dim RegCtrl As Boolean
Dim strcap As String
Dim fileini As String
Dim idfile As String
Dim strproc As String
Dim strview As String
Public setKoneksi As Integer

Private Sub Form_Load()
    
    
    Dim hwnd As Long
    
    
    
    
    strcap = "Transfer Data SID"
        
    
    fileini = "SIDANS.ini"
    idfile = FindFirstFile("C:\WINDOWS\SYSTEM\" & fileini, finfo)
    If idfile = -1 Then
       MsgBox "File konfigurasi tidak ditemukan", vbInformation, "File Config Tidak Ada !"
       Konfigurasi.pemanggil = 1
       Konfigurasi.Show vbModal
'       Unload Me
       setKoneksi = 0
    Else
       
        server = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "server")
        user = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "user")
        port = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "port")
        driver = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "driver")
        siduser = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "siduser")
        dbs = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "database")
        dbpwd = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "password")
        usersql = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "dbsid", "userdbsid")
        pwdsql = ReadIniValue("C:\WINDOWS\SYSTEM\SIDANS.ini", "dbsid", "dbsidpwd")
        
        Call bukasid
        Call buka_koneksi
        'Call data_standar
        setKoneksi = 1
        
    End If
    
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(Me.hwnd, RGB(255, 255, 255), 200, LWA_ALPHA Or LWA_COLORKEY)
End Sub




Private Sub Timer1_Timer()
    Dim rsdatacek As New ADODB.Recordset
    Dim bprsandi As String
    Dim strcekdata As String
    Dim cek As String
    
    If setKoneksi = 0 Then
        Unload Me
        Exit Sub
    End If
    
    Static j As Integer
    
    If j < 100 Then
        j = j + 1
    End If
    
    pb1.Value = j
    
    If idfile = -1 Then
        cek = cek
    Else
    
        If pb1.Value = 10 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa kolom dinrequest di tabel datapokok ...."
             strproc = "CREATE PROCEDURE addcol()" & vbCrLf & _
            "BEGIN" & vbCrLf & _
            "IF NOT EXISTS( SELECT * FROM information_schema.COLUMNS" & vbCrLf & _
            "WHERE COLUMN_NAME='dinrequest' AND TABLE_NAME='datapokok' AND TABLE_SCHEMA='" & dbs & "'" & vbCrLf & _
            ")" & vbCrLf & _
            "then" & vbCrLf & _
            "ALTER TABLE `" & dbs & "`.`datapokok` ADD COLUMN `dinrequest` INT(1) UNSIGNED NOT NULL DEFAULT 0;" & vbCrLf & _
            "END IF;" & vbCrLf & _
            "END;"
            dbbank.Execute ("drop procedure if exists addcol")
            dbbank.Execute (strproc)
            dbbank.Execute ("CALL addcol();")
            dbbank.Execute ("DROP PROCEDURE addcol;")
        
        ElseIf pb1.Value = 20 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa kolom din di tabel datapokok ...."
            strproc = "CREATE PROCEDURE addcol()" & vbCrLf & _
            "BEGIN" & vbCrLf & _
            "IF NOT EXISTS( SELECT * FROM information_schema.COLUMNS" & vbCrLf & _
            "WHERE COLUMN_NAME='din' AND TABLE_NAME='rekredit' AND TABLE_SCHEMA='" & dbs & "'" & vbCrLf & _
            ")" & vbCrLf & _
            "then" & vbCrLf & _
            "ALTER TABLE `" & dbs & "`.`rekredit` ADD COLUMN `din` VARCHAR(30) NOT NULL DEFAULT '-';" & vbCrLf & _
            "END IF;" & vbCrLf & _
            "END;"
            
            dbbank.Execute ("drop procedure if exists addcol")
            dbbank.Execute (strproc)
            dbbank.Execute ("CALL addcol();")
            dbbank.Execute ("DROP PROCEDURE addcol;")
        
        ElseIf pb1.Value = 30 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa id debitur  di tabel rekredit ...."
            strproc = "CREATE PROCEDURE addcol()" & vbCrLf & _
            "BEGIN" & vbCrLf & _
            "IF NOT EXISTS( SELECT * FROM information_schema.COLUMNS" & vbCrLf & _
            "WHERE COLUMN_NAME='cif_bi' AND TABLE_NAME='rekredit' AND TABLE_SCHEMA='" & dbs & "'" & vbCrLf & _
            ")" & vbCrLf & _
            "then" & vbCrLf & _
            "ALTER TABLE `" & dbs & "`.`rekredit` ADD COLUMN `cif_bi` VARCHAR(43) NOT NULL DEFAULT '-';" & vbCrLf & _
            "END IF;" & vbCrLf & _
            "END;"
            dbbank.Execute ("drop procedure if exists addcol")
            dbbank.Execute (strproc)
            dbbank.Execute ("CALL addcol();")
            dbbank.Execute ("DROP PROCEDURE addcol;")
            
        ElseIf pb1.Value = 40 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa kolom no_akad_awal di tabel rekredit ...."
            
             strproc = "CREATE PROCEDURE addcol()" & vbCrLf & _
            "BEGIN" & vbCrLf & _
            "IF NOT EXISTS( SELECT * FROM information_schema.COLUMNS" & vbCrLf & _
            "WHERE COLUMN_NAME='no_akad_awal' AND TABLE_NAME='rekredit' AND TABLE_SCHEMA='" & dbs & "'" & vbCrLf & _
            ")" & vbCrLf & _
            "then" & vbCrLf & _
            "ALTER TABLE `" & dbs & "`.`rekredit` ADD COLUMN `no_akad_awal` VARCHAR(25) NOT NULL DEFAULT '-';" & vbCrLf & _
            "END IF;" & vbCrLf & _
            "END;"
            dbbank.Execute ("drop procedure if exists addcol")
            dbbank.Execute (strproc)
            dbbank.Execute ("CALL addcol();")
            dbbank.Execute ("DROP PROCEDURE addcol;")
            
        ElseIf pb1.Value = 50 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa kolom no_akad_akhir di tabel rekredit ...."
            
            strproc = "CREATE PROCEDURE addcol()" & vbCrLf & _
            "BEGIN" & vbCrLf & _
            "IF NOT EXISTS( SELECT * FROM information_schema.COLUMNS" & vbCrLf & _
            "WHERE COLUMN_NAME='no_akad_akhir' AND TABLE_NAME='rekredit' AND TABLE_SCHEMA='" & dbs & "'" & vbCrLf & _
            ")" & vbCrLf & _
            "then" & vbCrLf & _
            "ALTER TABLE `" & dbs & "`.`rekredit` ADD COLUMN `no_akad_akhir` VARCHAR(25)  NOT NULL DEFAULT '-';" & vbCrLf & _
            "END IF;" & vbCrLf & _
            "END;"
            dbbank.Execute ("drop procedure if exists addcol")
            dbbank.Execute (strproc)
            dbbank.Execute ("CALL addcol();")
            dbbank.Execute ("DROP PROCEDURE addcol;")
            
        ElseIf pb1.Value = 60 Then
            
            Label1.Refresh
            Label1.Caption = "Memeriksa id_fasilitas pada rekbank ....."
            strproc = "CREATE PROCEDURE addcol()" & vbCrLf & _
            "BEGIN" & vbCrLf & _
            "IF NOT EXISTS( SELECT * FROM information_schema.COLUMNS" & vbCrLf & _
            "WHERE COLUMN_NAME='bi_fasilitas' AND TABLE_NAME='rekbank' AND TABLE_SCHEMA='" & dbs & "'" & vbCrLf & _
            ")" & vbCrLf & _
            "then" & vbCrLf & _
            "ALTER TABLE `" & dbs & "`.`rekbank` ADD COLUMN `bi_fasilitas` VARCHAR(40) NOT NULL DEFAULT '-';" & vbCrLf & _
            "END IF;" & vbCrLf & _
            "END;"
            dbbank.Execute ("drop procedure if exists addcol")
            dbbank.Execute (strproc)
            dbbank.Execute ("CALL addcol();")
            dbbank.Execute ("DROP PROCEDURE addcol;")
            
        ElseIf pb1.Value = 70 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa kolom suku_bunga di tabel rekbank ...."
            strproc = "CREATE PROCEDURE addcol()" & vbCrLf & _
            "BEGIN" & vbCrLf & _
            "IF NOT EXISTS( SELECT * FROM information_schema.COLUMNS" & vbCrLf & _
            "WHERE COLUMN_NAME='suku_bunga' AND TABLE_NAME='rekbank' AND TABLE_SCHEMA='" & dbs & "'" & vbCrLf & _
            ")" & vbCrLf & _
            "then" & vbCrLf & _
            "ALTER TABLE `" & dbs & "`.`rekbank` ADD COLUMN `suku_bunga` double DEFAULT 0;" & vbCrLf & _
            "END IF;" & vbCrLf & _
            "END;"
            dbbank.Execute ("drop procedure if exists addcol")
            dbbank.Execute (strproc)
            dbbank.Execute ("CALL addcol();")
            dbbank.Execute ("DROP PROCEDURE addcol;")
            
        ElseIf pb1.Value = 80 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa view data kredit ...."
            dbbank.Execute ("DROP VIEW IF EXISTS `data_kredit`")
            
            strview = "CREATE VIEW data_kredit AS (select rekredit.norek as norek, datapokok.din as DIN, rekredit.cif_bi as cif_bi, " & _
            "rekredit.tglmasuk as tgl_akad_awal, rekredit.tgljt as tgl_akad_akhir, rekredit.tglmasuk as tgl_awal_kredit, rekredit.tglmasuk as tgl_mulai, " & _
            "rekredit.tgljt as tgl_jt, rekredit.kali as kali, rekredit.guna as guna, rekredit.sektor as sektor, datapokok.idlokasi as idlokasi, " & _
            "bngkredit.pro1 as pro1, rekredit.carahitung as carahitung, rekredit.plafon as plafon, rekredit.gp as gp, rekredit.coll as coll, " & _
            "rekredit.ppap as ppap, rekredit.tglr as tglr from ((rekredit join bngkredit on ((rekredit.kodebng = bngkredit.kodebng))) join datapokok on ((datapokok.idnama = rekredit.idnama)))" & _
            " where (rekredit.saldoenc <> 0))"
            
            dbbank.Execute (strview)
        
        ElseIf pb1.Value = 85 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa kolom dinrequest di tabel rekredit ...."
             strproc = "CREATE PROCEDURE addcol()" & vbCrLf & _
            "BEGIN" & vbCrLf & _
            "IF NOT EXISTS( SELECT * FROM information_schema.COLUMNS" & vbCrLf & _
            "WHERE COLUMN_NAME='dinrequest' AND TABLE_NAME='rekredit' AND TABLE_SCHEMA='" & dbs & "'" & vbCrLf & _
            ")" & vbCrLf & _
            "then" & vbCrLf & _
            "ALTER TABLE `" & dbs & "`.`rekredit` ADD COLUMN `dinrequest` INT(1) UNSIGNED NOT NULL DEFAULT 0;" & vbCrLf & _
            "END IF;" & vbCrLf & _
            "END;"
            dbbank.Execute ("drop procedure if exists addcol")
            dbbank.Execute (strproc)
            dbbank.Execute ("CALL addcol();")
            dbbank.Execute ("DROP PROCEDURE addcol;")
        
        ElseIf pb1.Value = 90 Then
        
            Label1.Refresh
            Label1.Caption = "Memeriksa view xtunggakan ...."
            dbbank.Execute ("DROP VIEW IF EXISTS `qr_xtgk`")
            strview = "CREATE VIEW qr_xtgk AS ( SELECT `angsuran`.`Norek` AS `Norek`, MAX(`angsuran`.`TP`) AS `XTgkPokok`, MAX(`angsuran`.`TB`) AS `XTgkBunga` FROM `angsuran` GROUP BY `angsuran`.`Norek`)"
            dbbank.Execute (strview)
            
            Label1.Refresh
            Label1.Caption = "Memeriksa data bpr ...."
            
            
            
            strcekdata = "select * from idbpr"
            Set rsdatacek = dbbank.Execute(strcekdata)
            
            If Len(rsdatacek!sandibpr) < 5 Then
                bprsandi = InputBox("Masukkan sandi bpr: ", "Sandi BPR")
                dbbank.Execute ("update idbpr set sandibpr = '" & bprsandi & "'")
            End If
            
        End If
             
        
        If pb1.Value = 100 Then
            Label1.Refresh
            Label1.Caption = "Siap digunakan ...."
            Call tutup
        End If
    End If
End Sub


Private Function tutup()
    
    Unload Me
    main.Show
End Function

