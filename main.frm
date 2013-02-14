VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form main 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Transfer Data SID"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12870
   FillColor       =   &H00FFFFFF&
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6915
   ScaleWidth      =   12870
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      Picture         =   "main.frx":058A
      ScaleHeight     =   1545
      ScaleWidth      =   1545
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer trmText 
      Interval        =   100
      Left            =   12240
      Top             =   6000
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   6000
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6540
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   8811
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4080
      Width           =   12855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Width           =   12855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone : +62 21 515 5720 Fax : +62 21 515 5725"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jakarta 12190"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jl. Jend. Sudirman Kav. 52-53"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Equity Tower Lantai 26F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PT. ARTHA NUSA SEMBADA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "SISTEM TRANSFER DATA SID "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   1800
      Width           =   12855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   12855
   End
   Begin VB.Menu atur 
      Caption         =   "Pengaturan"
      Begin VB.Menu koneksi 
         Caption         =   "koneksi Database"
      End
      Begin VB.Menu penempatan 
         Caption         =   "Mapping Penempatan"
      End
      Begin VB.Menu conf 
         Caption         =   "Sinkronisasi"
      End
   End
   Begin VB.Menu export 
      Caption         =   "Ekspor Data"
      Begin VB.Menu din 
         Caption         =   "DIN"
      End
      Begin VB.Menu sid 
         Caption         =   "Data SID"
      End
   End
   Begin VB.Menu utility 
      Caption         =   "Utility"
      Begin VB.Menu menuvalidasi 
         Caption         =   "Periksa Data hasil Validasi"
      End
      Begin VB.Menu cekbaki 
         Caption         =   "Cek Baki Debet"
      End
      Begin VB.Menu calc 
         Caption         =   "Kalkulator"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public namalengkap As String
Dim strcap As String


Private Sub calc_Click()
    Shell ("calc.exe")
End Sub

Private Sub cekbaki_Click()
    compare.Show 1
End Sub

Private Sub conf_Click()
    sinc.Show 1
End Sub

Private Sub din_Click()
    Form1.Show 1
End Sub

Public Function data_standar()
    On Error GoTo salah
    
    Dim rsusersid As ADODB.Recordset
    
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
    
    strusersid = "select namalengkap from t_sys_users where userid = '" & siduser & "'"
    Set rsusersid = New ADODB.Recordset
    Set rsusersid = dbsid.Execute(strusersid)
    
    If rsusersid.EOF Then
        Unload main
        Konfigurasi.Show 1
        Load main
    Else
        namalengkap = rsusersid!namalengkap
    
    End If
    Label7.Width = Screen.Width
    Label8.Width = Screen.Width
    Label7.Caption = rsid!namabpr
    Label8.Caption = rsid!alamat
    Exit Function
salah:
    MsgBox ("Username, atau Password koneksi database salah")
    Konfigurasi.Show 1
    
End Function

Private Sub Form_Load()
     
    Dim r As RECT
    Dim hwnd As Long
    
    strcap = "Transfer Data SID"
        
    
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
        
        Call bukasid
        Call buka_koneksi
        Call data_standar
        
    End If
    
    

    hwnd = FindWindow("Shell_traywnd", vbNullString)

    GetWindowRect hwnd, r

    Me.Height = Screen.Height - ((r.Bottom - r.Top) * Screen.TwipsPerPixelX)
    Me.Width = Screen.Width
    
    'StatusBar1.Panels(2).Text = Format(Now, "dd")
    
    Shape1.Width = Me.Width
    Shape2.Width = Me.Width
    
    nama_bulan = MonthName(Format(Now, "mm"), True)
    
    StatusBar1.Panels(2).Text = Hari(Format(Now, "dd")) & " , " & Format(Now, "dd") & " " & nama_bulan & " " & Format(Now, "yyyy")
    
    StatusBar1.Panels(3).Text = "User SID : " & namalengkap
    
End Sub

Private Sub koneksi_Click()
     Konfigurasi.Show 1
End Sub



Private Sub menuvalidasi_Click()
    editValidasi.Show 1
End Sub

Private Sub penempatan_Click()
    frmrekbank.Show 1
End Sub

Private Sub sid_Click()
    frmDataSid.Show 1
End Sub

Private Sub Timer1_Timer()
    Select Case FormatJam
    Case DuabelasJam
        StatusBar1.Panels(1).Text = Format(Time, "hh:mm:ss AM/PM")
    Case Else
        StatusBar1.Panels(1).Text = Format(Time, "hh:mm:ss")
End Select
End Sub


Public Function Hari(ByVal Tanggal As Date) As String
Select Case Weekday(Tanggal)
Case 1: Hari = "Minggu"
Case 2: Hari = "Senin"
Case 3: Hari = "Selasa"
Case 4: Hari = "Rabu"
Case 5: Hari = "Kamis"
Case 6: Hari = "Jum'at"
Case 7: Hari = "Sabtu"
End Select
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub trmText_Timer()
   If Label1.Caption <> strcap Then
     If Label1.Alignment = 0 Then
        'run from left
        Label1.Caption = Left(strcap, Len(Label1.Caption) + 1)
     ElseIf Label1.Alignment = 1 Then
        'run from right
        Label1.Caption = Right(strcap, Len(Label1.Caption) + 1)
     ElseIf lblCaption.Alignment = 2 Then
        'run from the middle
       Label1.Caption = Mid(strcap, Len(strcap) \ 2 + Len(strcap) Mod 2 - Num, _
                            2 * (Num + 1) - Len(strcap) Mod 2)
       Num = Num + 1
     End If
   Else
      Label1.Caption = ""
      Num = 0
   End If
End Sub
