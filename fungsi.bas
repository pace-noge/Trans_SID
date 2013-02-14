Attribute VB_Name = "fungsi"
Public jns_data As String
Public id_data As String
Public id_bank As String
Public id_kntr_cab As String
Public create_date As String
Public create_user As String
Public ttl_record As String
Public regist As String
Public do_aksi As String
Public data_din As String
Public din As String
Public nama As String
Public alias As String
Public almt As String
Public kopos As String
Public kel As String
Public kec As String
Public dati2 As String
Public status As String
Public ket_status As String
Public npwp As String
Public ktp As String
Public tmpt_lahir As String
Public tgl_lahir As String
Public ibu As String
Public jenkel As String
Public paspor As String
Public jenis As String
Public operation As String
Public cif As String
Dim header As String
Dim footer As String
Dim isi As String
Public rsdata As ADODB.Recordset
Public strdata As String
Public rsdataDebitur As ADODB.Recordset
Public rsDataKredit As ADODB.Recordset
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Public Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public vPathSVC As String
Public vPathSVC1 As String
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * 260
  cAlternate As String * 14
End Type

Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Global finfo As WIN32_FIND_DATA

Public Function FormatText(StrFieldVal As String) As String
    Dim ChrPos As Long, PosFound As Long
    Dim WrkStr As String
    For ChrPos = 1 To Len(StrFieldVal)
        PosFound = InStr(ChrPos, StrFieldVal, "'")
        If PosFound > 0 Then
            WrkStr = WrkStr & Mid(StrFieldVal, ChrPos, PosFound - ChrPos + 1) & "'"
            ChrPos = PosFound
        Else
            WrkStr = WrkStr & Mid(StrFieldVal, ChrPos, Len(StrFieldVal))
            ChrPos = Len(StrFieldVal)
        End If
    Next ChrPos
    FormatText = WrkStr
End Function

Public Function TranslucentForm(Frm As Form, TranslucenceLevel As Byte) As Boolean
    SetWindowLong Frm.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Frm.hwnd, 0, TranslucenceLevel, LWA_ALPHA
    TranslucentForm = Err.LastDllError = 0
End Function


