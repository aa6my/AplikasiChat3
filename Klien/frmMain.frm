VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Klien"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsvBarang 
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9763
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   285
      Left            =   3855
      TabIndex        =   9
      Top             =   120
      Width           =   960
   End
   Begin VB.CheckBox chkCekStok 
      Caption         =   "Stok"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   930
      Width           =   1095
   End
   Begin VB.TextBox txtStok2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Text            =   "0"
      Top             =   930
      Width           =   615
   End
   Begin VB.TextBox txtStok1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "0"
      Top             =   930
      Width           =   615
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtNamaBarang 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "mie"
      Top             =   525
      Width           =   3495
   End
   Begin VB.CommandButton cmdCekBarang 
      Caption         =   "Cek Barang"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4935
      TabIndex        =   0
      Top             =   525
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   5400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "s.d"
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   930
      Width           =   210
   End
   Begin VB.Label Label2 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Private Const LOCAL_PORT    As Long = 1007

Private Const REC_SPR       As String * 1 = "|" 'separator baris
Private Const FLD_SPR       As String * 1 = "#" 'separator kolom

Dim tmp                     As String
Dim packageHdr              As String

Private Function rep(ByVal Kata As String) As String
    rep = Replace(Kata, "'", "''")
End Function

Private Function getQueryBarang(ByVal namaBarang As String, Optional ByVal stok1 As Integer = 0, Optional ByVal stok2 As Integer = 0) As String
    If chkCekStok.Value = vbChecked Then
        getQueryBarang = "WHERE nama LIKE '%" & rep(namaBarang) & "%' AND stok BETWEEN " & stok1 & " AND " & stok2 & ""
    Else
        getQueryBarang = "WHERE nama LIKE '%" & rep(namaBarang) & "%'"
    End If
End Function

Private Sub execOutput(ByVal data As String)
    Dim rec()   As String
    Dim fld()   As String
    
    Dim x       As Long
    Dim noUrut  As Long
    
    On Error GoTo errHandle
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    If Left(data, 2) = "~~" Then 'complete
        data = Replace(data, "~~", "")

    ElseIf data = "EOF" Then
        'do nothing
        
    Else
        data = Left(data, Len(data) - 1) 'remove ~ left
        data = Right(data, Len(data) - 1) 'remove ~ right
    End If

    lsvBarang.ListItems.Clear
    
    If data = "EOF" Then
        Screen.MousePointer = vbDefault
        MsgBox "Data barang dengan keyword '" & txtNamaBarang.Text & "' tidak ditemukan", vbInformation, "Informasi"
        
    Else
        'contoh data :
        '~~SUSU KEDELAI ABC 200M#1000#24|SUSU KEDELAI MELILEA 500#1000#0|KOPI SUSU KPL API 3P#1000#0|SUSU KEDELAI ABC 200#1000#2
        '| -> pemisah baris
        '# -> pemisah kolom
        
        rec = Split(data, REC_SPR)
        
        With lsvBarang
            noUrut = 1
            For x = LBound(rec) To UBound(rec)
                fld = Split(rec(x), FLD_SPR)
                
                .ListItems.Add , , noUrut
                .ListItems(noUrut).SubItems(1) = fld(0) 'nama barang
                .ListItems(noUrut).SubItems(2) = FormatNumber(fld(1), 0) 'harga
                .ListItems(noUrut).SubItems(3) = fld(2) 'stok
                
                noUrut = noUrut + 1
            Next x
        End With
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandle:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation, "Warning"
End Sub

Private Function startConnect(ByVal ipServer As String) As Boolean
    On Error Resume Next
    
    If Socket.State <> sckClosed Then Socket.Close ' close existing connection
    Call Socket.Connect(ipServer, LOCAL_PORT)
    With Socket
        Do While .State <> sckConnected
            DoEvents
            If .State = sckError Then Exit Function
        Loop
    End With
    
    startConnect = True
End Function

Private Function send(ByVal strData As String) As Boolean
    If Socket.State = sckConnected Then
        Call Socket.SendData(strData)
        DoEvents
        
    Else
        send = False
        Exit Function
    End If
   
    send = True
End Function

Private Sub chkCekStok_Click()
    txtStok1.Enabled = CBool(chkCekStok.Value)
    txtStok2.Enabled = txtStok1.Enabled
End Sub

Private Sub cmdConnect_Click()
    lsvBarang.ListItems.Clear
    cmdCekBarang.Enabled = False
    
    If cmdConnect.Caption = "Connect" Then
        If startConnect(txtServer.Text) Then
            cmdConnect.Caption = "Disconnet"
            cmdCekBarang.Enabled = True
        End If
    Else
        Socket.Close
        cmdConnect.Caption = "Connect"
    End If
End Sub

Private Sub cmdCekBarang_Click()
    Dim param   As String
    
    If Not (Len(txtNamaBarang.Text) >= 3) Then
        MsgBox "Minimal nama barang 3 huruf", vbExclamation, "Warning"
        txtNamaBarang.SetFocus
        
        Exit Sub
    End If
    
    tmp = ""
    packageHdr = ""
    
    If chkCekStok.Value = vbChecked Then
        param = getQueryBarang(txtNamaBarang.Text, Val(txtStok1.Text), Val(txtStok2.Text))
    Else
        param = getQueryBarang(txtNamaBarang.Text)
    End If
    
    If Not send(param) Then MsgBox "Mengirim data gagal", vbExclamation, "Warning"
End Sub

Private Sub Form_Load()
    With lsvBarang
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        
        .ColumnHeaders.Add , , "No.", 500
        .ColumnHeaders.Add , , "Nama Barang", 3000
        .ColumnHeaders.Add , , "Harga", 1350, lvwColumnRight
        .ColumnHeaders.Add , , "Stok", 700, lvwColumnRight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Socket.State <> sckClosed Then Socket.Close        ' close existing connection
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim dataMasuk   As String
    
    'On Error Resume Next
    
    Socket.GetData dataMasuk
    
    If Left(dataMasuk, 2) = "~~" Then 'package data <= 1024
        Call execOutput(dataMasuk)
    
    ElseIf dataMasuk = "EOF" Then 'data tidak ditemukan
        Call execOutput(dataMasuk)
        
    Else
        'package data > 1024
        'berikut kode untuk penggabungan package data
        tmp = tmp & dataMasuk
        If InStr(1, dataMasuk, "~") > 0 Then packageHdr = packageHdr & "~"
        
        If Len(packageHdr) = 2 Then Call execOutput(tmp) 'penggabungan package data selesai
    End If
End Sub

