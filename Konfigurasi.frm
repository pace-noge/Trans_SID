VERSION 5.00
Begin VB.Form Konfigurasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Tes Koneksi"
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtpwddbsid 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   20
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox txtuserdbsid 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      TabIndex        =   18
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtport 
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Top             =   2040
      Width           =   2655
   End
   Begin VB.ComboBox ListDriverODBC 
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      Top             =   600
      Width           =   2655
   End
   Begin VB.ComboBox cbodatabase 
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txtusersid 
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtpwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtuser 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtserver 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label11 
      Caption         =   "Password SQL Server"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Username SQL Server"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "*) Wajib di isi sesuai username SID"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label Label8 
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Port Database ANS"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Driver Database ANS"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Username Pelapor SID"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Database ANS"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Password Database ANS"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "User Database ANS"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Alamat Server ANS"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Konfigurasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pemanggil As Integer


Private Sub cbodatabase_DropDown()
    
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
       driver = ListDriverODBC.Text
       server = txtserver.Text
       user = txtuser.Text
      dbpwd = txtpwd.Text
       'NamaDatabase = txtdb.Text
       NamaDatabase = ""
       port = txtport.Text
     Call buka_koneksi
Strsql = "SELECT SCHEMA_NAME AS `Database` From INFORMATION_SCHEMA.SCHEMATA order by SCHEMA_NAME asc"
Set rs = dbbank.Execute(Strsql)
    If Not rs.EOF Then
        cbodatabase.Clear
        While Not rs.EOF
        
            cbodatabase.AddItem rs(0)
        rs.MoveNext
        Wend
    End If
End Sub

Private Sub cmdCancel_Click()
    If pemanggil = 1 Then
        Unload Me
        splash.Show
    Else
        Unload Me
        
    End If
End Sub

Private Sub cmdsave_Click()
    
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "server", txtserver.Text
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "user", txtuser.Text
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "port", txtport.Text
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "driver", ListDriverODBC.Text
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "siduser", txtusersid.Text
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "database", cbodatabase.Text
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "SIDEXIM", "password", txtpwd.Text
    
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "dbsid", "userdbsid", txtuserdbsid.Text
    WriteIniValue "C:\WINDOWS\SYSTEM\SIDANS.ini", "dbsid", "dbsidpwd", txtpwddbsid.Text
    
    server = txtserver.Text
    user = txtuser.Text
    port = txtport.Text
    driver = ListDriverODBC.Text
    siduser = txtusersid.Text
    dbs = cbodatabase.Text
    dbpwd = txtpwd.Text
    usersql = txtuserdbsid.Text
    pwdsql = txtpwddbsid.Text
    
    MsgBox ("Pengaturan telah di simpan"), vbInformation, "Simpan Konfigurasi"
    
End Sub



Private Sub Command1_Click()
        driver = ListDriverODBC.Text
       server = txtserver.Text
       user = txtuser.Text
       'NamaUser = Decrypt(NamaUser, True)
       'UserPwd = ReadINI(FILEini, "User_Pwd", "C:\WINDOWS\SYSTEM\" & FILEini)
       dbpwd = txtpwd.Text
       dbs = cbodatabase.Text
       port = txtport.Text
       'Komp_server = txtKomServer.Text
       
     
       
Call buka_koneksi
    If statuskonek = 1 Then
        MsgBox "Koneksi Dengan Database BPR Sukses", vbInformation + vbOKOnly, "Info"
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
   ' Fill a listbox control with the list of all available DSNs
    Dim odbcTool As New odbcTool.Dsn
    Dim Dsn() As String, i As Long
    
    If odbcTool.GetOdbcDriverList(Dsn()) Then
    ' a True return value means success
    ListDriverODBC.Clear
    For i = 0 To UBound(Dsn)
        ListDriverODBC.AddItem Dsn(i)
    Next
    Else
    ' a False value means error
    MsgBox "Tidak dapat membaca daftar DSN", vbExclamation
    End If
    
    fileini = "SIDANS.ini"
    idfile = FindFirstFile("C:\WINDOWS\SYSTEM\" & fileini, finfo)
    If idfile = -1 Then
       'MsgBox ("File konfigurasi tidak ditemukan"), vbInformation, "Info"
    Else
       'path_database = ReadINI(FILEini, "DATABASE PATH", "C:\WINDOWS\SYSTEM\" & FILEini)\
              
       ListDriverODBC.Text = driver
       txtport.Text = port
       txtpwd.Text = dbpwd
       txtuser.Text = user
       cbodatabase.AddItem dbs
       cbodatabase.Text = dbs
       txtserver.Text = server
       txtusersid = siduser
       txtuserdbsid.Text = usersql
       txtpwddbsid.Text = pwdsql
       
      
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    
    For i = Me.Left To (Screen.Width / 2) Step 10
    Me.Height = Me.Height - 15
    Me.Width = Me.Width - 20
    Me.Left = Me.Left + 100
    DoEvents
    Next
    Unload Me
End Sub


