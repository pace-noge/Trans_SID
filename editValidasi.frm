VERSION 5.00
Begin VB.Form editValidasi 
   Caption         =   "Edit Data Validasi"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   2745
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Periksa Tunggakan"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
      Begin VB.CommandButton cektgk 
         Caption         =   "cek"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "No Rekening"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "editValidasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cektgk_Click()
    Dim rscektgk As ADODB.Recordset
    Dim strcektgk As String
    
    strcektgk = "SELECT rekredit.plafon, rekredit.saldoenc, angsuran.* FROM angsuran, rekredit WHERE angsuran.norek = rekredit.norek and angsuran.norek = '" & Text1.Text & "' AND STATUS = 'Tertunggak' ORDER BY angsuran"
    Set rscektgk = New ADODB.Recordset
    Set rscektgk = dbbank.Execute(strcektgk)
    
    If Not rscektgk.EOF Or Not rscektgk.BOF Then
        Text2.Text = "Tunggakan pokok    : " & rscektgk!bpokok & vbNewLine & "Tunggakan Bunga    : " & rscektgk!bbunga & vbNewLine & "Tanggal Tunggakan : " & Format(rscektgk!tgl, "dd-mm-yyyy") & vbNewLine & "Plafon                        : " & rscektgk!plafon & vbNewLine & "Baki Debet                : " & rscektgk!saldoenc
    End If
    
End Sub
