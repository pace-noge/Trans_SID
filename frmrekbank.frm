VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmrekbank 
   Caption         =   "Data Penempatan"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15675
   LinkTopic       =   "Form2"
   ScaleHeight     =   7815
   ScaleWidth      =   15675
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgsid 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   4048
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
   Begin VB.Frame Frame1 
      Caption         =   "Data Penempatan SID"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   15015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Rekbank ANS"
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   15015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4683
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
Attribute VB_Name = "frmrekbank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub BuatKolom()
With MSFlexGrid1
.Cols = 7
.Rows = 20
.Col = 0
.Row = 0
.Text = "No Rekening"
.ColWidth(0) = 1000
.CellAlignment = flexAlignCenterCenter
.Col = 1
.Row = 0
.Text = "Nama Nasabah"
.ColWidth(1) = 3000
.CellAlignment = flexAlignCenterCenter
.Col = 2
.Row = 0
.Text = "saldo"
.ColWidth(2) = 1000
.CellAlignment = flexAlignCenterCenter
.Col = 3
.Row = 0
.Text = "Tgl Jatuh Tempo"
.ColWidth(3) = 2000
.CellAlignment = flexAlignCenterCenter
.Col = 4
.Row = 0
.Text = "Jenis"
.ColWidth(4) = 10
.CellAlignment = flexAlignCenterCenter
.Col = 5
.Row = 0
.Text = "ID BI Fasilitas"
.ColWidth(4) = 2000
.CellAlignment = flexAlignCenterCenter
Col = 6
.Row = 0
.Text = "Suku bunga"
.ColWidth(4) = 2000
.CellAlignment = flexAlignCenterCenter
.TextMatrix(1, 0) = 1
End With
End Sub

Sub TampilData()
On Error Resume Next
Dim BarisData As Integer
Dim rs As ADODB.Recordset
MSFlexGrid1.Clear
BuatKolom
MSFlexGrid1.Rows = 4
BarisData = 0
Set rs = New ADODB.Recordset
Dim Perintah As String
Set rs.ActiveConnection = dbbank
rs.CursorLocation = adUseClient
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
Perintah = "SELECT rekbank.norek, datapokok.nama, rekbank.saldoenc, rekbank.tgljt, rekbank.jenis, rekbank.bi_fasilitas FROM datapokok, rekbank WHERE rekbank.idnama = datapokok.idnama AND rekbank.jenis IN ('10', '20', '30')  ORDER BY norek"
rs.Open (Perintah)
Set DataGrid1.DataSource = rs

DataGrid1.Columns(0).Width = Len(DataGrid1.Columns(0).Text) * 100
DataGrid1.Columns(1).Width = Len(DataGrid1.Columns(1).Text) * 200
DataGrid1.Columns(2).Width = Len(DataGrid1.Columns(2).Text) * 150
DataGrid1.Columns(3).Width = Len(DataGrid1.Columns(3).Text) * 100
DataGrid1.Columns(5).Width = Len(DataGrid1.Columns(5).Text) * 120


End Sub

Private Sub Form_Load()
    Dim rssyncrekbank As ADODB.Recordset
    
    Dim rsrekbankans As ADODB.Recordset
    Dim str As String
    
    'str = "SELECT rekbank.norek, datapokok.nama, rekbank.saldoenc, rekbank.tgljt AS `jatuh tempo`, rekbank.jenis, rekbank.bi_fasilitas FROM datapokok, rekbank WHERE rekbank.idnama = datapokok.idnama AND saldoenc <> 0"
    str = "SELECT T_PENEMPATAN.ID_FASILITAS, T_PENEMPATAN.JW_BULAN, T_PENEMPATAN.KOLEKTIBILITAS, T_PENEMPATAN.SUKU_BUNGA, T_PENEMPATAN.KETERANGAN , REF_BANK.DESC2, T_PENEMPATAN.NILAI_PENEMPATAN FROM T_PENEMPATAN INNER JOIN REF_BANK ON T_PENEMPATAN.SANDI_BANK = REF_BANK.DESC1"
    Set rssyncrekbank = New ADODB.Recordset
    Set rssyncrekbank = dbsid.Execute(str)
    
    Set dgsid.DataSource = rssyncrekbank
    dgsid.Columns(0).Width = Len(dgsid.Columns(0).Text) * 100
    dgsid.Columns(5).Width = 2800
    
    
    TampilData
    
End Sub


Private Sub MSFlexGrid1_Click()
    With MSFlexGrid1
        txtDataEntry.Text = .TextMatrix(.Row, .Col)
        txtDataEntry.Move .CellLeft + .Left, .CellTop + .Top, _
          .CellWidth, .CellHeight
        txtDataEntry.Visible = True
        DoEvents
        txtDataEntry.SetFocus
        .TextMatrix(.Row, .Col) = txtDataEntry.Text
    End With
 
End Sub
