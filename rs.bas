Attribute VB_Name = "rs"
Public dbbank As New ADODB.Connection
Public dbsid As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rstgl As New ADODB.Recordset
Public rsid As ADODB.Recordset
Global idbank As String
Global idkancab As String
Global rsttlkred As ADODB.Recordset
Global bln As String
Global thn As String
Global idlembaga As String
Global server As String
Global user As String
Global port As String
Global dbs As String
Global siduser As String
Global driver As String
Global dbpwd As String
Global statsid As Integer
Global usersql As String
Global pwdsql As String
Global statuskonek As Integer

Public Function buka_koneksi()
    
    NamaDriver = driver
    Alamatserver = server
    NamaDatabase = dbs
    NamaUser = user
    UserPwd = dbpwd
    No_port = port
    
    On Error GoTo salah
'CARA -> 1  UNTUK MEMBUKA DATA ACCES BISA JUGA DENGAN CARA 1 INI
''wdbfolder = App.Path
'--------------------------------------------------------------------------
  
        Set dbbank = New ADODB.Connection
        dbbank.CursorLocation = adUseClient
    With dbbank
  '  strcon = "Driver={sql server};server=wawan;database=BPR;uid=aku;pwd=login;"
  ' strcon = "Driver=mysql odbc 5.1 driver;server=localhost;database=BPR;uid=root;pwd=kosongin;port=3306"
   strcon = "driver=" & NamaDriver & ";server=" & Alamatserver & ";database=" & NamaDatabase & ";uid=" & NamaUser & ";pwd=" & UserPwd & ";port=" & No_port & ""
    
        .CursorLocation = adUseClient
          If .State = adStateOpen Then
            .Close
            .Open (strcon)
            Else
            .Open (strcon)
            End If
    End With
    statuskonek = 1
'--------------------------------------------------------------------------
Exit Function
salah:
statuskonek = 0
MsgBox Err.Description & " Gagal Koneksi Database " & Err.Number, vbCritical + vbOKOnly, "Konek Database"
       
End Function

Sub bukasid()
    Dim nmrError As String
    Dim ketError As String

    On Error GoTo salah
    
    Set dbsid = New ADODB.Connection
    dbsid.CursorLocation = adUseClient
    
    With dbsid
        .Provider = "SQLOLEDB"
        .Properties("Data Source") = ".\sqlexpress"
        .Properties("User ID") = usersql
        .Properties("password") = pwdsql
        .Open
        .DefaultDatabase = "Sidbpr"
    End With
    
    If dbsid.State = adStateOpen Then
        statsid = 1
    Else
        bra = bh
    End If
        
    Exit Sub
    
salah:
    nmrError = Err.Number
    ketError = Err.Description
    statsid = 0
    MsgBox nmrError & "Harus di komputer yang terinstall database sid" & ketError & ", gagal Konek"
        
        
End Sub

