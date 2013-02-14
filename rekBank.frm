VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form rekBank 
   Caption         =   "Form2"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "Form2"
   ScaleHeight     =   6990
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   5295
      Left            =   7200
      TabIndex        =   1
      Top             =   480
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6855
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4695
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8281
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "rekBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsExample As ADODB.Recordset
Dim str As String

Private Sub Form_Load()
    Dim lngWidth As Long
Dim intLoopCount As Integer
Const SCROLL_BAR_WIDTH = 320

With MSFlexGrid1
    
    str = "select idnama, norek, saldoenc from rekredit"
    Set rsExample = New ADODB.Recordset
    Set rsExample = dbbank.Execute(str)
    
    With rsExample
    .MoveLast
    'set number of grid rows and columns
    'to fit recordset
    MSFlexGrid1.Rows = .RecordCount
    MSFlexGrid1.Cols = 3
    .MoveFirst

    Do
        If !nama < 0 Then
            'reference cell to set cell text colour
            'to Red for shareprice
            MSFlexGrid1.Row = .AbsolutePosition + 1
            MSFlexGrid1.Col = 1
            MSFlexGrid1.CellForeColor = vbRed
            MSFlexGrid1.TextMatrix(.AbsolutePosition + 1, 1) = _
              !shareprice
            MSFlexGrid1.TextMatrix(.AbsolutePosition + 1, 2) = _
              !movement
        ElseIf !movement = 0 Then
            'reference cell to set cell background
            'colour to Yellow for shareprice
            MSFlexGrid1.Row = .AbsolutePosition + 1
            MSFlexGrid1.Col = 1
            MSFlexGrid1.CellBackColor = vbYellow
            MSFlexGrid1.TextMatrix(.AbsolutePosition + 1, 1) = _
              !shareprice
            MSFlexGrid.TextMatrix(.AbsolutePosition + 1, 2) = _
              !movement
        Else
            MSFlexGrid1.TextMatrix(.AbsolutePosition + 1, 1) = _
              !shareprice
            MSFlexGrid1.TextMatrix(.AbsolutePosition + 1, 2) = _
              !movement
        End If
        .MoveNext
    Loop Until .EOF
    End With
End With

End Sub
