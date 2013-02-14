VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form compare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validasi Baki Debet SID dan BPR"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   5760
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtselisih 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar pgbar 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid fgbakidebet 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8281
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      FillStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox bakisid 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox bakibpr 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Caption         =   "Cek Baki Debet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3840
         MaskColor       =   &H00C0C000&
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   3720
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Selisih"
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
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Baki Debet SID"
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Baki Debet BPR"
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
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "compare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    On Error Resume Next
    
    Dim rsbakidebetans As ADODB.Recordset
    Dim rsbakidebetsid As ADODB.Recordset
    Dim no As Integer
    Dim totsaldoenc As Double
    Dim totbakidebet As Double
    Dim selisih As Double
    
    totsaldoenc = 0
    
    no = 1
    
    strbakidebetsid = "select no_rekening, baki_debet from t_kredit where len(no_rekening) > 7 order by no_rekening"
    Set rsbakidebetsid = New ADODB.Recordset
    Set rsbakidebetsid = dbsid.Execute(strbakidebetsid)
    
    strbakidebetans = "select norek, saldoenc from rekredit"
    Set rsbakidebetans = New ADODB.Recordset
    Set rsbakidebetans = dbbank.Execute(strbakidebetans)
    
    
    pgbar.Max = rsbakidebetsid.RecordCount
    pgbar.Value = 0
    With fgbakidebet
        .Rows = rsbakidebetsid.RecordCount
        .Cols = 3
        
        .TextMatrix(0, 0) = "NO Rekening"
        .TextMatrix(0, 1) = "Baki Debet SID"
        .TextMatrix(0, 2) = "Baki Debet BPR"
        .ColWidth(0) = 1300
        .ColWidth(1) = 1800
        .ColWidth(2) = 1800
        While Not rsbakidebetsid.EOF
        
            pgbar.Value = pgbar.Value + 1
            
            strbakidebetans = "select saldoenc from rekredit where norek = '" & rsbakidebetsid!no_rekening & "'"
            Set rsbakidebetans = New ADODB.Recordset
            Set rsbakidebetans = dbbank.Execute(strbakidebetans)
            
            If Not rsbakidebetans.BOF Then
            
            .Redraw = True
             .FillStyle = flexFillSingle
            If Val(rsbakidebetsid!baki_debet) = Val(rsbakidebetans!saldoenc) Then
                .ForeColor = vbBlack
                .BackColor = vbWhite
                
            Else
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End If
            
            .TextMatrix(no, 0) = rsbakidebetsid!no_rekening
            .TextMatrix(no, 1) = rsbakidebetsid!baki_debet
            .TextMatrix(no, 2) = rsbakidebetans!saldoenc
            
            
            no = no + 1
            
            totsaldoenc = totsaldoenc + rsbakidebetans!saldoenc
            totbakidebet = totbakidebet + rsbakidebetsid!baki_debet
            
        rsbakidebetsid.MoveNext
        
        Else
        
            strbakidebetans = "select norek from logkredit where norek = '" & rsbakidebetsid!no_rekening & "'"
            Set rsbakidebetans = New ADODB.Recordset
            Set rsbakidebetans = dbbank.Execute(strbakidebetans)
            .Redraw = True
             .FillStyle = flexFillSingle
            If Val(rsbakidebetsid!baki_debet) = Val(rsbakidebetans!saldoenc) Then
                .ForeColor = vbBlack
                .BackColor = vbWhite
                
            Else
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End If
            
            .TextMatrix(no, 0) = rsbakidebetsid!no_rekening
            .TextMatrix(no, 1) = rsbakidebetsid!baki_debet
            .TextMatrix(no, 2) = "0"
            
            
            no = no + 1
            
            totsaldoenc = totsaldoenc + rsbakidebetans!saldoenc
            totbakidebet = totbakidebet + rsbakidebetsid!baki_debet
        
        rsbakidebetsid.MoveNext
        End If
        Wend
        
        
        
        
    End With
    
    bakibpr.Text = totsaldoenc
    bakisid.Text = totbakidebet
    txtselisih.Text = totbakidebet - totsaldoenc
    
    If txtselisih.Text <> 0 Then
        txtselisih.BackColor = vbRed
        txtselisih.ForeColor = vbWhite
    Else
        txtselisih.BackColor = vbWhite
        txtselisih.ForeColor = vbBlack
    End If
        
    
'    bakibpr.Locked
'    bakisid.Locked
    
End Sub

Private Sub Form_Load()
    pgbar.Value = 0
End Sub
