VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form data_agunan 
   Caption         =   "Data Agunan"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form2"
   ScaleHeight     =   9135
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Debitur 
      Caption         =   "Debitur"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11895
      Begin VB.CommandButton Command4 
         Caption         =   ">>"
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox tbxktp 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox tbxalmt 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox tbxnama 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox tbxidnama 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "NPWP"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "No. KTP"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "No. Rekening"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Input Data Aguanan"
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   11895
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   2760
         TabIndex        =   48
         Text            =   "-"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   ">>"
         Height          =   375
         Left            =   4800
         TabIndex        =   47
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">"
         Height          =   375
         Left            =   4440
         TabIndex        =   46
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<"
         Height          =   375
         Left            =   4080
         TabIndex        =   45
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<<"
         Height          =   375
         Left            =   3720
         TabIndex        =   44
         Top             =   3840
         Width           =   375
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   2760
         TabIndex        =   43
         Text            =   "-"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2760
         TabIndex        =   41
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2760
         TabIndex        =   39
         Top             =   2880
         Width           =   4695
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2760
         TabIndex        =   37
         Top             =   2640
         Width           =   4695
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2760
         TabIndex        =   35
         Top             =   2400
         Width           =   4695
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2760
         TabIndex        =   33
         Text            =   "-"
         Top             =   1800
         Width           =   4695
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2760
         TabIndex        =   31
         Top             =   2160
         Width           =   4695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2760
         TabIndex        =   29
         Top             =   1560
         Width           =   7575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2760
         TabIndex        =   27
         Top             =   1320
         Width           =   6015
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2760
         TabIndex        =   25
         Top             =   1080
         Width           =   6015
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2760
         TabIndex        =   23
         Text            =   "-"
         Top             =   840
         Width           =   3735
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3600
         TabIndex        =   21
         Text            =   "-"
         Top             =   480
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3600
         TabIndex        =   19
         Text            =   "-"
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label19 
         Caption         =   "Di Asuransikan"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "Paripasu (%)"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "Tanggal Penilaian"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "Nilai Agunan (Penilai Independent)"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label14 
         Caption         =   "Nilai Agunan (Bank)"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Dati II Lokasi Agunan"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "Niai Agunan (NJOP)"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Status / Bukti Kepemilikan"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Nama Pemilik Agunan"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Peringkat Surat Berharga"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Peringkat Surat Berharga"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Jenis Agunan"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agunan"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   11895
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   2566
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "data_agunan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dataagunanans As New ADODB.Recordset
Dim dataAgunan As New ADODB.Recordset
Dim posisidata  As Double
Dim ttldata As Double
Dim rspengikatan As ADODB.Recordset
Dim rsjenisJaminan As ADODB.Recordset
Dim rsasuransi As ADODB.Recordset
Dim rsPilihItem As String












Private Sub Command1_Click()
    strdataagunanans = "SELECT rekredit.idnama, rekredit.norek, datapokok.BUKTIDIRI,rekredit.pengikatan, datapokok.alamat, datapokok.nama, rekredit.JAMINAN, datapokok.IDLOKASI, rekredit.NILAIJAMINAN, rekredit.NILAIJAMINAN1, rekredit.TGLMASUK FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama order by norek limit 1, 1"
    Set dataagunanans = New ADODB.Recordset
    Set dataagunanans = dbbank.Execute(strdataagunanans)
    
    tbxidnama.Text = dataagunanans!norek
    tbxnama.Text = dataagunanans!nama
    tbxalmt.Text = dataagunanans!alamat
    tbxktp.Text = dataagunanans!buktidiri
End Sub

Private Sub Command2_Click()
    
    If IsEmpty(posisidata) Then
        posisidata = 1
    ElseIf posisidata = 0 Then
        posisidata = 1
    Else
        posisidata = posisidata - 1
    End If
    
    ambilData (posisidata)
End Sub

Private Sub Command3_Click()
    
    If IsEmpty(posisidata) Then
        posisidata = 0
    End If
    
    posisidata = posisidata + 1
    
    ambilData (posisidata)
    
    


End Sub

Private Sub Command4_Click()
    strdataagunanans = "SELECT rekredit.idnama, rekredit.norek, datapokok.BUKTIDIRI, datapokok.alamat, datapokok.nama FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama order by norek"
    Set dataagunanans = New ADODB.Recordset
    Set dataagunanans = dbbank.Execute(strdataagunanans)
    
    posisidata = dataagunanans.RecordCount - 1
    
    ambilData (posisidata)
    
    
End Sub

Private Sub Form_Load()
    
    Dim rsjnisagunan As ADODB.Recordset
    
    strdataagunanans = "SELECT rekredit.idnama, rekredit.norek, datapokok.BUKTIDIRI, datapokok.alamat, datapokok.nama FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama order by norek"
    Set dataagunanans = New ADODB.Recordset
    Set dataagunanans = dbbank.Execute(strdataagunanans)
    
    ttldata = dataagunanans.RecordCount
    
    tbxidnama.Text = dataagunanans!norek
    tbxnama.Text = dataagunanans!nama
    tbxalmt.Text = dataagunanans!alamat
    tbxktp.Text = dataagunanans!buktidiri
    
    strpengikatan = "select * from ref_jenis_pengikatan"
    Set rspengikatan = New ADODB.Recordset
    Set rspengikatan = dbbank.Execute(strpengikatan)
    
    While Not rspengikatan.EOF
    
        Combo5.AddItem rspengikatan!desc1
        Combo2.AddItem rspengikatan!desc2
        
        
    rspengikatan.MoveNext
    Wend
    
    strjnisagunan = "select * from ref_jenis_agunan"
    Set rsjnisagunan = New ADODB.Recordset
    Set rsjnisagunan = dbbank.Execute(strjnisagunan)
    
    While Not rsjnisagunan.EOF
        
        Combo1.AddItem rsjnisagunan!desc2
        rsjnisagunan.MoveNext
    Wend
    
End Sub


'Private Sub Combo5_Change()
'    tambahItem(combo2.Text, "desc2", "desc1", "ref_jenis_pngikatan", "combo2", "desc2")
'End Sub
'
'Private Function tambahItem(desc As String, descsumber As String, descTarget As String, tabel As String, namaCombo As String, descDiminta As String)
'
'    strpilih = "select " & descDiminta & " from " & tabel & " where descSumber = '" & desc & "'"
'    Set rsPilihItem = New ADODB.Recordset
'    Set rspilih = dbbank.Execute(strpilih)
'
'    If namaCombo = "combo1" Then
'        Combo1.SelText = rspilih!descDiminta
'    ElseIf namaCombo = "Combo2" Then
'        Combo2.SelText = rspilih!descDiminta
'    ElseIf namaCombo = "combo3" Then
'        Combo3.SelText = rspilih!descDiminta
'    ElseIf namaCombo = "combo4" Then
'        Combo4.SelText = rspilih!descDiminta
'    ElseIf namaCombo = "combo5" Then
'        Combo5.SelText = rspilih!descDiminta
'    End If
'
'
'
'End Function

Private Sub tbxidnama_Change()

    strdataagunan = "SELECT * from data_agunan WHERE norek =  '" & dataagunanans!norek & "'"
    Set datagunan = New ADODB.Recordset
    Set dataAgunan = dbbank.Execute(strdataagunan)
    
    Set DataGrid1.DataSource = dataAgunan
    
    
    
    
End Sub


Private Function ambilData(Posisi As Double)
    
    
    strdataagunanans = "SELECT rekredit.idnama, rekredit.norek, datapokok.BUKTIDIRI, datapokok.alamat, datapokok.nama FROM datapokok, rekredit WHERE datapokok.idnama = rekredit.idnama order by norek limit " & Posisi & ", 1"
    Set dataagunanans = New ADODB.Recordset
    Set dataagunanans = dbbank.Execute(strdataagunanans)
    
    If Not dataagunanans.BOF Then
        tbxidnama.Text = dataagunanans!norek
        tbxnama.Text = dataagunanans!nama
        tbxalmt.Text = dataagunanans!alamat
        tbxktp.Text = dataagunanans!buktidiri
    Else
        MsgBox ("Ini adalah Posisi data terakhir")
    End If
End Function
